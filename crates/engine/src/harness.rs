//! Test harness for workbook operations with event tracking.
//!
//! This module provides `EngineHarness`, a wrapper around `Workbook` that:
//! - Tracks events (BatchApplied, CellsChanged, RevisionChanged)
//! - Tracks undo groups
//! - Provides `apply_ops` with atomic/non-atomic semantics
//!
//! Use this harness to test session server invariants without GUI dependencies.

use std::cell::RefCell;
use std::rc::Rc;

use crate::cell_id::CellId;
use crate::events::{
    BatchAppliedEvent, BatchError, CellsChangedEvent, EventCollector, RevisionChangedEvent,
    WorkbookEvent,
};
use crate::workbook::Workbook;

/// Operation to apply to a workbook.
#[derive(Debug, Clone)]
pub enum Op {
    /// Set a cell's value (auto-detects formulas).
    SetCellValue {
        sheet_index: usize,
        row: usize,
        col: usize,
        value: String,
    },
    /// Set a cell's formula explicitly.
    SetCellFormula {
        sheet_index: usize,
        row: usize,
        col: usize,
        formula: String,
    },
    /// Clear a cell.
    ClearCell {
        sheet_index: usize,
        row: usize,
        col: usize,
    },
    /// Simulate an error at this op (for testing rollback).
    #[cfg(test)]
    SimulateError { message: String },
}

/// Result of applying operations.
#[derive(Debug, Clone)]
pub struct ApplyResult {
    /// Number of ops successfully applied.
    pub applied: usize,
    /// Revision after apply (new on success, unchanged on full rollback).
    pub revision: u64,
    /// Error if any op failed.
    pub error: Option<BatchError>,
}

/// Minimal undo group tracker.
#[derive(Debug, Default)]
pub struct UndoTracker {
    /// Number of undo groups created.
    group_count: usize,
    /// Current group depth (for nested groups).
    depth: usize,
    /// Whether the current group was aborted.
    aborted: bool,
}

impl UndoTracker {
    pub fn new() -> Self {
        Self::default()
    }

    /// Begin a new undo group.
    pub fn begin_group(&mut self) {
        self.depth += 1;
    }

    /// End the current undo group. If not aborted, increments group count.
    pub fn end_group(&mut self) {
        if self.depth > 0 {
            self.depth -= 1;
            if self.depth == 0 && !self.aborted {
                self.group_count += 1;
            }
            if self.depth == 0 {
                self.aborted = false;
            }
        }
    }

    /// Abort the current group (rollback). No undo entry is created.
    pub fn abort_group(&mut self) {
        self.aborted = true;
        if self.depth > 0 {
            self.depth -= 1;
        }
    }

    /// Returns the number of committed undo groups.
    pub fn group_count(&self) -> usize {
        self.group_count
    }

    /// Returns current nesting depth.
    pub fn depth(&self) -> usize {
        self.depth
    }
}

/// Test harness wrapping Workbook with event and undo tracking.
pub struct EngineHarness {
    workbook: Workbook,
    events: Rc<RefCell<EventCollector>>,
    undo: UndoTracker,
}

impl EngineHarness {
    /// Create a new harness with a fresh workbook.
    pub fn new() -> Self {
        Self {
            workbook: Workbook::new(),
            events: Rc::new(RefCell::new(EventCollector::new())),
            undo: UndoTracker::new(),
        }
    }

    /// Create a harness wrapping an existing workbook.
    pub fn with_workbook(workbook: Workbook) -> Self {
        Self {
            workbook,
            events: Rc::new(RefCell::new(EventCollector::new())),
            undo: UndoTracker::new(),
        }
    }

    /// Get a reference to the underlying workbook.
    pub fn workbook(&self) -> &Workbook {
        &self.workbook
    }

    /// Get a mutable reference to the underlying workbook.
    pub fn workbook_mut(&mut self) -> &mut Workbook {
        &mut self.workbook
    }

    /// Get collected events.
    pub fn events(&self) -> std::cell::Ref<'_, EventCollector> {
        self.events.borrow()
    }

    /// Clear collected events.
    pub fn clear_events(&self) {
        self.events.borrow_mut().clear();
    }

    /// Get undo group count (committed groups only).
    pub fn undo_group_count(&self) -> usize {
        self.undo.group_count()
    }

    /// Current revision.
    pub fn revision(&self) -> u64 {
        self.workbook.revision()
    }

    /// Apply operations with specified atomicity.
    ///
    /// - `atomic=true`: All-or-nothing. On error, rollback all changes.
    /// - `atomic=false`: Partial apply. On error, keep changes up to failing op.
    ///
    /// Events emitted:
    /// - On success: RevisionChanged, CellsChanged, BatchApplied
    /// - On atomic rollback: BatchApplied (with error, no other events)
    /// - On partial failure: RevisionChanged, CellsChanged (for applied ops), BatchApplied
    pub fn apply_ops(&mut self, ops: &[Op], atomic: bool) -> ApplyResult {
        let prev_revision = self.workbook.revision();
        let total_ops = ops.len();

        // Begin undo group
        self.undo.begin_group();

        // Track cells changed during this batch
        let mut changed_cells: Vec<CellId> = Vec::new();

        // Begin batch (defers recalc)
        self.workbook.begin_batch();

        let mut applied = 0;
        let mut error: Option<BatchError> = None;

        for (idx, op) in ops.iter().enumerate() {
            match self.apply_single_op(op, &mut changed_cells) {
                Ok(()) => {
                    applied += 1;
                }
                Err(e) => {
                    error = Some(BatchError {
                        code: e.code,
                        message: e.message,
                        op_index: idx,
                    });
                    break;
                }
            }
        }

        if error.is_some() && atomic {
            // Atomic rollback: revert all changes
            self.rollback_batch(&changed_cells);
            self.undo.abort_group();

            // Emit only BatchApplied (no CellsChanged, no RevisionChanged)
            let result = ApplyResult {
                applied: 0,
                revision: prev_revision,
                error: error.clone(),
            };

            self.events.borrow_mut().push(WorkbookEvent::BatchApplied(
                BatchAppliedEvent {
                    revision: prev_revision,
                    applied: 0,
                    total: total_ops,
                    error,
                },
            ));

            return result;
        }

        // End batch (triggers recalc)
        // Note: end_batch increments revision if changes were made
        self.workbook.end_batch();

        // End undo group (only if not aborted)
        self.undo.end_group();

        let new_revision = self.workbook.revision();

        // Emit events for successful or partial apply
        if new_revision != prev_revision {
            // RevisionChanged
            self.events
                .borrow_mut()
                .push(WorkbookEvent::RevisionChanged(RevisionChangedEvent {
                    revision: new_revision,
                    previous: prev_revision,
                }));

            // CellsChanged (only for this revision)
            if !changed_cells.is_empty() {
                self.events
                    .borrow_mut()
                    .push(WorkbookEvent::CellsChanged(CellsChangedEvent {
                        revision: new_revision,
                        cells: changed_cells,
                    }));
            }
        }

        // BatchApplied
        self.events
            .borrow_mut()
            .push(WorkbookEvent::BatchApplied(BatchAppliedEvent {
                revision: new_revision,
                applied,
                total: total_ops,
                error: error.clone(),
            }));

        ApplyResult {
            applied,
            revision: new_revision,
            error,
        }
    }

    /// Apply a single op, tracking changed cells.
    fn apply_single_op(
        &mut self,
        op: &Op,
        changed_cells: &mut Vec<CellId>,
    ) -> Result<(), OpError> {
        match op {
            Op::SetCellValue {
                sheet_index,
                row,
                col,
                value,
            } => {
                let sheet_id = self
                    .workbook
                    .sheet_id_at_idx(*sheet_index)
                    .ok_or_else(|| OpError {
                        code: "invalid_sheet".to_string(),
                        message: format!("Sheet index {} not found", sheet_index),
                    })?;

                self.workbook
                    .set_cell_value_tracked(*sheet_index, *row, *col, value);
                changed_cells.push(CellId::new(sheet_id, *row, *col));
                Ok(())
            }
            Op::SetCellFormula {
                sheet_index,
                row,
                col,
                formula,
            } => {
                let sheet_id = self
                    .workbook
                    .sheet_id_at_idx(*sheet_index)
                    .ok_or_else(|| OpError {
                        code: "invalid_sheet".to_string(),
                        message: format!("Sheet index {} not found", sheet_index),
                    })?;

                // SetCellFormula uses the same path as SetCellValue
                // (formulas are auto-detected by the = prefix)
                self.workbook
                    .set_cell_value_tracked(*sheet_index, *row, *col, formula);
                changed_cells.push(CellId::new(sheet_id, *row, *col));
                Ok(())
            }
            Op::ClearCell {
                sheet_index,
                row,
                col,
            } => {
                let sheet_id = self
                    .workbook
                    .sheet_id_at_idx(*sheet_index)
                    .ok_or_else(|| OpError {
                        code: "invalid_sheet".to_string(),
                        message: format!("Sheet index {} not found", sheet_index),
                    })?;

                self.workbook.clear_cell_tracked(*sheet_index, *row, *col);
                changed_cells.push(CellId::new(sheet_id, *row, *col));
                Ok(())
            }
            #[cfg(test)]
            Op::SimulateError { message } => Err(OpError {
                code: "simulated_error".to_string(),
                message: message.clone(),
            }),
        }
    }

    /// Rollback changes made during a batch (for atomic mode).
    fn rollback_batch(&mut self, changed_cells: &[CellId]) {
        // Clear the batch_changed list to prevent recalc
        self.workbook.batch_changed.clear();
        // Manually decrement batch depth to exit batch mode without recalc
        if self.workbook.batch_depth > 0 {
            self.workbook.batch_depth -= 1;
        }

        // In a real implementation, we'd restore previous values here.
        // For now, we just clear the cells (which is lossy but sufficient
        // for testing the event/revision invariants).
        for cell_id in changed_cells {
            if let Some(sheet_idx) = self.workbook.sheet_index_by_id(cell_id.sheet) {
                if let Some(sheet) = self.workbook.sheet_mut(sheet_idx) {
                    sheet.clear_cell(cell_id.row, cell_id.col);
                }
            }
        }
    }
}

impl Default for EngineHarness {
    fn default() -> Self {
        Self::new()
    }
}

/// Internal error type for op application.
struct OpError {
    code: String,
    message: String,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_harness_basic_apply() {
        let mut harness = EngineHarness::new();

        let ops = vec![
            Op::SetCellValue {
                sheet_index: 0,
                row: 0,
                col: 0,
                value: "hello".to_string(),
            },
            Op::SetCellValue {
                sheet_index: 0,
                row: 0,
                col: 1,
                value: "world".to_string(),
            },
        ];

        let result = harness.apply_ops(&ops, false);

        assert_eq!(result.applied, 2);
        assert!(result.error.is_none());
        assert_eq!(harness.revision(), 1);

        // Check events
        let events = harness.events();
        assert_eq!(events.revision_changed().len(), 1);
        assert_eq!(events.cells_changed().len(), 1);
        assert_eq!(events.batch_applied().len(), 1);

        // CellsChanged should have both cells
        let cells_changed = &events.cells_changed()[0];
        assert_eq!(cells_changed.revision, 1);
        assert_eq!(cells_changed.cells.len(), 2);
    }

    #[test]
    fn test_harness_atomic_rollback() {
        let mut harness = EngineHarness::new();

        let ops = vec![
            Op::SetCellValue {
                sheet_index: 0,
                row: 0,
                col: 0,
                value: "first".to_string(),
            },
            Op::SimulateError {
                message: "test error".to_string(),
            },
            Op::SetCellValue {
                sheet_index: 0,
                row: 0,
                col: 1,
                value: "never reached".to_string(),
            },
        ];

        let initial_rev = harness.revision();
        let result = harness.apply_ops(&ops, true); // atomic=true

        // Atomic rollback: nothing applied, revision unchanged
        assert_eq!(result.applied, 0);
        assert!(result.error.is_some());
        assert_eq!(result.revision, initial_rev);
        assert_eq!(harness.revision(), initial_rev);

        // Check events: only BatchApplied (no CellsChanged, no RevisionChanged)
        let events = harness.events();
        assert_eq!(events.revision_changed().len(), 0, "no RevisionChanged on rollback");
        assert_eq!(events.cells_changed().len(), 0, "no CellsChanged on rollback");
        assert_eq!(events.batch_applied().len(), 1);

        let batch = &events.batch_applied()[0];
        assert_eq!(batch.applied, 0);
        assert!(batch.error.is_some());
    }

    #[test]
    fn test_harness_partial_apply() {
        let mut harness = EngineHarness::new();

        let ops = vec![
            Op::SetCellValue {
                sheet_index: 0,
                row: 0,
                col: 0,
                value: "first".to_string(),
            },
            Op::SetCellValue {
                sheet_index: 0,
                row: 0,
                col: 1,
                value: "second".to_string(),
            },
            Op::SimulateError {
                message: "test error".to_string(),
            },
            Op::SetCellValue {
                sheet_index: 0,
                row: 0,
                col: 2,
                value: "never reached".to_string(),
            },
        ];

        let initial_rev = harness.revision();
        let result = harness.apply_ops(&ops, false); // atomic=false

        // Partial apply: 2 ops applied, revision incremented
        assert_eq!(result.applied, 2);
        assert!(result.error.is_some());
        assert_eq!(result.revision, initial_rev + 1);

        // Check events: RevisionChanged + CellsChanged + BatchApplied
        let events = harness.events();
        assert_eq!(events.revision_changed().len(), 1);
        assert_eq!(events.cells_changed().len(), 1);
        assert_eq!(events.batch_applied().len(), 1);

        // CellsChanged should have only the 2 applied cells
        let cells_changed = &events.cells_changed()[0];
        assert_eq!(cells_changed.cells.len(), 2);
    }

    #[test]
    fn test_harness_undo_tracking() {
        let mut harness = EngineHarness::new();

        // Successful batch creates undo group
        let ops = vec![Op::SetCellValue {
            sheet_index: 0,
            row: 0,
            col: 0,
            value: "test".to_string(),
        }];
        harness.apply_ops(&ops, false);
        assert_eq!(harness.undo_group_count(), 1);

        // Another successful batch
        harness.apply_ops(&ops, false);
        assert_eq!(harness.undo_group_count(), 2);

        // Atomic rollback does NOT create undo group
        harness.clear_events();
        let ops_with_error = vec![
            Op::SetCellValue {
                sheet_index: 0,
                row: 0,
                col: 0,
                value: "will rollback".to_string(),
            },
            Op::SimulateError {
                message: "fail".to_string(),
            },
        ];
        harness.apply_ops(&ops_with_error, true);
        assert_eq!(harness.undo_group_count(), 2, "rollback should not add undo group");
    }
}
