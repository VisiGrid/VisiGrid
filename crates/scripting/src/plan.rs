//! Adapter from Lua's private journal to the producer-neutral engine plan.

use visigrid_engine::operation_plan::{
    CellCoordinate, CellRange, ExecutionContextFingerprint, GroupId, OperationGroup,
    OperationMetadata, PlannedCellValue, PlannedOp, PlannedOperation, ProducerClaim,
    VerificationDefinition,
};
use visigrid_engine::workbook::Workbook;

use crate::lua_formulas::published_functions_fingerprint;
use crate::{LuaCellValue, LuaOp, LuaReviewMetadata};

/// Producer-neutral pieces extracted from Lua's private journal.
#[derive(Debug, Clone)]
pub struct LuaPlanParts {
    pub operations: Vec<PlannedOperation>,
    pub groups: Vec<OperationGroup>,
    pub verification: Vec<VerificationDefinition>,
}

/// Convert the Lua journal into executable operations, untrusted review
/// descriptions, and engine-evaluated assertion definitions. Metadata markers
/// apply to subsequent mutations until another marker replaces or clears them.
pub fn lua_journal_to_plan(ops: &[LuaOp]) -> Result<LuaPlanParts, String> {
    let mut metadata = OperationMetadata::default();
    let mut groups: Vec<OperationGroup> = Vec::new();
    let mut planned = Vec::new();
    let mut verification = Vec::new();

    for operation in ops {
        if let LuaOp::SetReviewMetadata(review) = operation {
            metadata = operation_metadata(review);
            if let Some(group_id) = &review.group_id {
                if let Some(existing) = groups
                    .iter_mut()
                    .find(|group| group.id.0 == *group_id)
                {
                    if let Some(title) = &review.group_title {
                        if existing.title == *group_id {
                            existing.title = title.clone();
                        } else if existing.title != *title {
                            return Err(format!(
                                "review group '{group_id}' was declared with conflicting titles"
                            ));
                        }
                    }
                    if let Some(description) = &review.group_description {
                        match &existing.description {
                            None => existing.description = Some(description.clone()),
                            Some(current) if current == description => {}
                            Some(_) => return Err(format!(
                                "review group '{group_id}' was declared with conflicting descriptions"
                            )),
                        }
                    }
                } else {
                    groups.push(OperationGroup {
                        id: GroupId(group_id.clone()),
                        title: review.group_title.clone().unwrap_or_else(|| group_id.clone()),
                        description: review.group_description.clone(),
                    });
                }
            }
            continue;
        }
        if let LuaOp::RequestVerification(request) = operation {
            if request.kind != "gross_minus_group_equals_preview" {
                return Err(format!(
                    "unsupported verification kind '{}'",
                    request.kind
                ));
            }
            let ((start_row, start_col), (end_row, end_col)) =
                crate::ops::parse_range(&request.source_range).ok_or_else(|| {
                    format!("invalid verification source range '{}'", request.source_range)
                })?;
            let amount_column_name = request.amount_column.trim();
            if amount_column_name.is_empty()
                || !amount_column_name.bytes().all(|byte| byte.is_ascii_alphabetic())
            {
                return Err(format!(
                    "invalid verification amount column '{}'",
                    request.amount_column
                ));
            }
            let (_, amount_column) = crate::ops::parse_a1(&format!("{amount_column_name}1"))
            .ok_or_else(|| {
                format!(
                    "invalid verification amount column '{}'",
                    request.amount_column
                )
            })?;
            verification.push(VerificationDefinition::RetainedTotal {
                id: request.id.clone(),
                label: request.label.clone(),
                source_range: CellRange {
                    start: CellCoordinate {
                        row: start_row - 1,
                        col: start_col - 1,
                    },
                    end: CellCoordinate {
                        row: end_row - 1,
                        col: end_col - 1,
                    },
                },
                amount_column: amount_column - 1,
                excluded_group: GroupId(request.excluded_group.clone()),
                tolerance: request.tolerance,
                currency: request.currency.clone(),
            });
            continue;
        }

        if let Some(operation) = lua_op_to_planned_op(operation) {
            planned.push(PlannedOperation {
                operation,
                metadata: metadata.clone(),
            });
        }
    }

    Ok(LuaPlanParts {
        operations: planned,
        groups,
        verification,
    })
}

fn operation_metadata(review: &LuaReviewMetadata) -> OperationMetadata {
    OperationMetadata {
        group_id: review.group_id.clone().map(GroupId),
        reason: review.reason.clone().map(ProducerClaim),
        sources: review.sources.clone(),
    }
}

fn lua_op_to_planned_op(operation: &LuaOp) -> Option<PlannedOp> {
    Some(match operation {
        LuaOp::SetReviewMetadata(_) | LuaOp::RequestVerification(_) => return None,
        LuaOp::SetValue {
            row,
            col,
            value: LuaCellValue::Nil,
        } => PlannedOp::ClearCell {
            coordinate: CellCoordinate {
                row: *row as usize,
                col: *col as usize,
            },
        },
        LuaOp::SetValue { row, col, value } => PlannedOp::SetCellValue {
            coordinate: CellCoordinate {
                row: *row as usize,
                col: *col as usize,
            },
            value: match value {
                LuaCellValue::Nil => unreachable!("handled above"),
                LuaCellValue::Number(value) => PlannedCellValue::Number(*value),
                LuaCellValue::String(value) => PlannedCellValue::Text(value.clone()),
                LuaCellValue::Bool(value) => PlannedCellValue::Boolean(*value),
                LuaCellValue::Error(value) => PlannedCellValue::Error(value.clone()),
            },
        },
        LuaOp::SetFormula { row, col, formula } => PlannedOp::SetCellFormula {
            coordinate: CellCoordinate {
                row: *row as usize,
                col: *col as usize,
            },
            formula: formula.clone(),
        },
        LuaOp::ClearCell { row, col } => PlannedOp::ClearCell {
            coordinate: CellCoordinate {
                row: *row as usize,
                col: *col as usize,
            },
        },
        LuaOp::DeleteRows { at, count } => PlannedOp::DeleteRows {
            at: *at as usize,
            count: *count as usize,
        },
        LuaOp::SetCellStyle {
            r1,
            c1,
            r2,
            c2,
            style,
        } => PlannedOp::SetCellStyle {
            range: CellRange {
                start: CellCoordinate {
                    row: *r1 as usize,
                    col: *c1 as usize,
                },
                end: CellCoordinate {
                    row: *r2 as usize,
                    col: *c2 as usize,
                },
            },
            style: visigrid_engine::cell::CellStyle::from_int(*style as i32),
        },
    })
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExecutionContextGenerationKey {
    pub functions_generation: u64,
    pub functions_source_hash: String,
    pub locale: String,
    pub timezone: String,
    pub auto_recalc: bool,
    pub iterative_calculation_enabled: bool,
    pub iterative_max_iterations: u32,
    pub iterative_tolerance_bits: u64,
}

/// Cheap key for deciding whether a cached execution-context check remains
/// valid. Formula volatility is intentionally absent: formula changes bump
/// workbook revision and trigger the complete context scan on the cache miss.
pub fn execution_context_generation_key(workbook: &Workbook) -> ExecutionContextGenerationKey {
    let functions = published_functions_fingerprint();
    let locale = std::env::var("LC_ALL")
        .or_else(|_| std::env::var("LC_NUMERIC"))
        .or_else(|_| std::env::var("LANG"))
        .unwrap_or_else(|_| "C".into());
    let timezone =
        std::env::var("TZ").unwrap_or_else(|_| chrono::Local::now().offset().to_string());
    ExecutionContextGenerationKey {
        functions_generation: functions.generation,
        functions_source_hash: functions.source_hash,
        locale,
        timezone,
        auto_recalc: workbook.auto_recalc(),
        iterative_calculation_enabled: workbook.iterative_enabled(),
        iterative_max_iterations: workbook.iterative_max_iters(),
        iterative_tolerance_bits: workbook.iterative_tolerance().to_bits(),
    }
}

/// Fingerprint calculation inputs that are not represented by workbook
/// revision. This accessor is intentionally read-only and never reloads Lua.
pub fn execution_context_fingerprint(
    workbook: &Workbook,
    planned_ops: &[PlannedOperation],
) -> ExecutionContextFingerprint {
    let generation = execution_context_generation_key(workbook);
    let mut volatile_inputs = Vec::new();
    for sheet in workbook.sheets() {
        for (_, cell) in sheet.cells_iter() {
            note_volatile_formula(&cell.value.raw_display(), &mut volatile_inputs);
        }
    }
    for planned in planned_ops {
        if let PlannedOp::SetCellFormula { formula, .. } = &planned.operation {
            note_volatile_formula(formula, &mut volatile_inputs);
        }
    }
    volatile_inputs.sort();
    volatile_inputs.dedup();

    ExecutionContextFingerprint {
        engine_version: env!("CARGO_PKG_VERSION").into(),
        functions_generation: generation.functions_generation,
        functions_source_hash: generation.functions_source_hash,
        locale: generation.locale,
        timezone: generation.timezone,
        auto_recalc: generation.auto_recalc,
        iterative_calculation_enabled: generation.iterative_calculation_enabled,
        iterative_max_iterations: generation.iterative_max_iterations,
        iterative_tolerance_bits: generation.iterative_tolerance_bits,
        volatile_inputs,
    }
}

fn note_volatile_formula(formula: &str, found: &mut Vec<String>) {
    if !formula.starts_with('=') {
        return;
    }
    let formula = formula.to_ascii_uppercase();
    for function in ["NOW", "TODAY", "RAND", "RANDBETWEEN"] {
        if formula.contains(&format!("{function}(")) {
            found.push(function.to_string());
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn lua_nil_and_explicit_clear_share_the_canonical_clear_operation() {
        let coordinate = CellCoordinate { row: 2, col: 3 };
        let parts = lua_journal_to_plan(&[
            LuaOp::SetValue {
                row: 2,
                col: 3,
                value: LuaCellValue::Nil,
            },
            LuaOp::ClearCell { row: 2, col: 3 },
        ])
        .unwrap();
        assert_eq!(
            parts
                .operations
                .iter()
                .map(|planned| &planned.operation)
                .collect::<Vec<_>>(),
            vec![
                &PlannedOp::ClearCell { coordinate },
                &PlannedOp::ClearCell { coordinate },
            ]
        );
    }

    #[test]
    fn review_metadata_is_associated_with_subsequent_operations() {
        let parts = lua_journal_to_plan(&[
            LuaOp::SetReviewMetadata(LuaReviewMetadata {
                group_id: Some("duplicates".into()),
                group_title: Some("Exact duplicates".into()),
                group_description: None,
                reason: Some("Same transaction identifier and amount".into()),
                sources: vec!["A2:C2".into(), "A4:C4".into()],
            }),
            LuaOp::DeleteRows { at: 3, count: 1 },
        ])
        .unwrap();

        assert_eq!(parts.groups[0].id, GroupId("duplicates".into()));
        assert!(parts.verification.is_empty());
        assert_eq!(
            parts.operations[0].metadata.group_id,
            Some(parts.groups[0].id.clone())
        );
        assert_eq!(parts.operations[0].metadata.sources.len(), 2);
    }

    #[test]
    fn execution_context_reports_volatile_formulas() {
        let mut workbook = Workbook::new();
        workbook.set_auto_recalc(false);
        workbook.set_iterative_enabled(true);
        workbook.set_iterative_max_iters(77);
        workbook.set_iterative_tolerance(0.000_123);
        workbook.set_cell_value_tracked(0, 0, 0, "=TODAY()");
        let context = execution_context_fingerprint(&workbook, &[]);
        assert!(context.volatile_inputs.contains(&"TODAY".to_string()));
        assert!(!context.functions_source_hash.is_empty());
        assert!(!context.auto_recalc);
        assert!(context.iterative_calculation_enabled);
        assert_eq!(context.iterative_max_iterations, 77);
        assert_eq!(context.iterative_tolerance_bits, 0.000_123_f64.to_bits());
    }

    #[test]
    fn transaction_fixture_materializes_verifies_applies_and_undoes() {
        use visigrid_engine::operation_plan::{
            ChangeKind, OperationPlanRequest, PlanId, PlanProducer, PreparedOperationPlan,
            VerificationDefinition, VerificationStatus,
        };

        let mut workbook = Workbook::new();
        for (row, values) in [
            (0, ["Transaction", "Vendor", "Amount"]),
            (1, ["tx-001", "Amazon.com", "100"]),
            (2, ["tx-002", "AMZN", "200"]),
            (3, ["tx-001", "Amazon.com", "100"]),
            (4, ["", "", ""]),
            (5, ["tx-003", "Acme", "50"]),
            (6, ["", "Total", "=SUM(C2:C4)"]),
        ] {
            for (col, value) in values.into_iter().enumerate() {
                if !value.is_empty() {
                    workbook.set_cell_value_tracked(0, row, col, value);
                }
            }
        }

        let runtime = crate::LuaRuntime::new().unwrap();
        let source = crate::SheetSnapshot::from_sheet(workbook.active_sheet());
        let result = runtime.eval_with_sheet(
            include_str!("../../../fixtures/review_mode/transaction_cleanup.lua"),
            Box::new(source),
        );
        assert!(result.error.is_none(), "fixture error: {:?}", result.error);
        let parts = lua_journal_to_plan(&result.ops).unwrap();
        let operations = parts.operations;
        let groups = parts.groups;
        let mut verification = parts.verification;
        verification.insert(
            0,
            VerificationDefinition::NoNewFormulaErrors {
                id: "no_new_errors".into(),
                label: Some("No new formula errors".into()),
            },
        );
        let context = execution_context_fingerprint(&workbook, &operations);
        let prepared = PreparedOperationPlan::materialize(
            &workbook,
            OperationPlanRequest {
                id: PlanId("pv_transaction_fixture".into()),
                workbook_session_id: "fixture-session".into(),
                source_sheet_id: workbook.active_sheet_id(),
                expected_revision: workbook.revision(),
                execution_context: context.clone(),
                producer: PlanProducer {
                    kind: "lua_fixture".into(),
                    name: "Transaction cleanup".into(),
                    source_path: Some("fixtures/review_mode/transaction_cleanup.lua".into()),
                    source_hash: None,
                },
                title: "Clean transaction export".into(),
                description: Some("Phase 1.5 deterministic dogfood fixture".into()),
                operations,
                groups,
                verification,
            },
        )
        .unwrap();

        assert_eq!(prepared.plan().groups.len(), 4);
        assert!(prepared.plan().changes.iter().any(|change| {
            change.kind == ChangeKind::RowDeleted
                && change.group_id == Some(GroupId("exact_duplicates".into()))
                && change.reason.is_some()
                && change.sources == ["A2:C2", "A4:C4"]
        }));
        assert_eq!(
            prepared
                .plan()
                .verification
                .iter()
                .find(|check| check.id == "retained_payments")
                .unwrap()
                .status,
            VerificationStatus::Passed
        );
        assert_eq!(
            prepared.preview_workbook().active_sheet().get_display(4, 2),
            "350"
        );

        let commit = prepared.verify_candidate(&workbook, &context).unwrap();
        assert!(commit
            .verification
            .iter()
            .any(|result| result.id == "retained_payments"
                && result.status == VerificationStatus::Passed));
        commit.redo_into(&mut workbook);
        assert_eq!(workbook.active_sheet().get_display(4, 2), "350");
        commit.undo_into(&mut workbook);
        assert_eq!(workbook.active_sheet().get_display(3, 0), "tx-001");
        assert_eq!(workbook.active_sheet().get_display(6, 2), "400");
    }
}
