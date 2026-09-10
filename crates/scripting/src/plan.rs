//! Adapter from Lua's private journal to the producer-neutral engine plan.

use visigrid_engine::operation_plan::{
    CellCoordinate, CellRange, ExecutionContextFingerprint, PlannedCellValue, PlannedOp,
};
use visigrid_engine::workbook::Workbook;

use crate::lua_formulas::published_functions_fingerprint;
use crate::{LuaCellValue, LuaOp};

pub fn lua_ops_to_planned_ops(ops: &[LuaOp]) -> Vec<PlannedOp> {
    ops.iter()
        .map(|operation| match operation {
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
        .collect()
}

/// Fingerprint calculation inputs that are not represented by workbook
/// revision. This accessor is intentionally read-only and never reloads Lua.
pub fn execution_context_fingerprint(
    workbook: &Workbook,
    planned_ops: &[PlannedOp],
) -> ExecutionContextFingerprint {
    let functions = published_functions_fingerprint();
    let locale = std::env::var("LC_ALL")
        .or_else(|_| std::env::var("LC_NUMERIC"))
        .or_else(|_| std::env::var("LANG"))
        .unwrap_or_else(|_| "C".into());
    let timezone =
        std::env::var("TZ").unwrap_or_else(|_| chrono::Local::now().offset().to_string());
    let mut volatile_inputs = Vec::new();
    for sheet in workbook.sheets() {
        for (_, cell) in sheet.cells_iter() {
            note_volatile_formula(&cell.value.raw_display(), &mut volatile_inputs);
        }
    }
    for operation in planned_ops {
        if let PlannedOp::SetCellFormula { formula, .. } = operation {
            note_volatile_formula(formula, &mut volatile_inputs);
        }
    }
    volatile_inputs.sort();
    volatile_inputs.dedup();

    ExecutionContextFingerprint {
        engine_version: env!("CARGO_PKG_VERSION").into(),
        functions_generation: functions.generation,
        functions_source_hash: functions.source_hash,
        locale,
        timezone,
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
        let operations = lua_ops_to_planned_ops(&[
            LuaOp::SetValue {
                row: 2,
                col: 3,
                value: LuaCellValue::Nil,
            },
            LuaOp::ClearCell { row: 2, col: 3 },
        ]);
        assert_eq!(
            operations,
            vec![
                PlannedOp::ClearCell { coordinate },
                PlannedOp::ClearCell { coordinate },
            ]
        );
    }

    #[test]
    fn execution_context_reports_volatile_formulas() {
        let mut workbook = Workbook::new();
        workbook.set_cell_value_tracked(0, 0, 0, "=TODAY()");
        let context = execution_context_fingerprint(&workbook, &[]);
        assert!(context.volatile_inputs.contains(&"TODAY".to_string()));
        assert!(!context.functions_source_hash.is_empty());
    }
}
