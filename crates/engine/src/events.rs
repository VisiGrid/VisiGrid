//! Event types for workbook change notifications.
//!
//! These events enable the session server to notify clients of changes
//! without polling. They're also used by the test harness to verify
//! invariants about event ordering and revision boundaries.

use crate::cell_id::CellId;

/// Events emitted by Workbook during batch operations.
#[derive(Debug, Clone, PartialEq)]
pub enum WorkbookEvent {
    /// A batch of operations was applied (successfully or partially).
    BatchApplied(BatchAppliedEvent),

    /// Cells changed as a result of operations or recalc.
    /// Always tagged with the revision that produced the changes.
    CellsChanged(CellsChangedEvent),

    /// Revision number changed. Emitted exactly once per successful batch.
    RevisionChanged(RevisionChangedEvent),
}

/// Emitted after apply_ops completes (success or partial failure).
#[derive(Debug, Clone, PartialEq)]
pub struct BatchAppliedEvent {
    /// Revision after this batch (new revision on success, unchanged on full rollback).
    pub revision: u64,
    /// Number of ops successfully applied.
    pub applied: usize,
    /// Total number of ops in the batch.
    pub total: usize,
    /// Error code if batch failed (None = full success).
    pub error: Option<BatchError>,
}

/// Error information for failed batch operations.
#[derive(Debug, Clone, PartialEq)]
pub struct BatchError {
    /// Error code (e.g., "formula_parse_error", "revision_mismatch").
    pub code: String,
    /// Human-readable message.
    pub message: String,
    /// Index of the failing op (0-based).
    pub op_index: usize,
}

/// Emitted when cells change value (from ops or recalc).
#[derive(Debug, Clone, PartialEq)]
pub struct CellsChangedEvent {
    /// Revision that produced these changes.
    /// INVARIANT: All cells in this event belong to this single revision.
    pub revision: u64,
    /// Cells that changed (may include recalc dependents).
    pub cells: Vec<CellId>,
}

/// Emitted exactly once per successful batch.
#[derive(Debug, Clone, PartialEq)]
pub struct RevisionChangedEvent {
    /// The new revision number.
    pub revision: u64,
    /// The previous revision number.
    pub previous: u64,
}

/// Callback type for receiving workbook events.
pub type EventCallback = Box<dyn FnMut(WorkbookEvent) + Send>;

/// Simple event collector for testing.
#[derive(Default)]
pub struct EventCollector {
    events: Vec<WorkbookEvent>,
}

impl EventCollector {
    pub fn new() -> Self {
        Self { events: Vec::new() }
    }

    pub fn push(&mut self, event: WorkbookEvent) {
        self.events.push(event);
    }

    pub fn events(&self) -> &[WorkbookEvent] {
        &self.events
    }

    pub fn clear(&mut self) {
        self.events.clear();
    }

    pub fn len(&self) -> usize {
        self.events.len()
    }

    pub fn is_empty(&self) -> bool {
        self.events.is_empty()
    }

    /// Filter to only BatchApplied events.
    pub fn batch_applied(&self) -> Vec<&BatchAppliedEvent> {
        self.events
            .iter()
            .filter_map(|e| match e {
                WorkbookEvent::BatchApplied(b) => Some(b),
                _ => None,
            })
            .collect()
    }

    /// Filter to only CellsChanged events.
    pub fn cells_changed(&self) -> Vec<&CellsChangedEvent> {
        self.events
            .iter()
            .filter_map(|e| match e {
                WorkbookEvent::CellsChanged(c) => Some(c),
                _ => None,
            })
            .collect()
    }

    /// Filter to only RevisionChanged events.
    pub fn revision_changed(&self) -> Vec<&RevisionChangedEvent> {
        self.events
            .iter()
            .filter_map(|e| match e {
                WorkbookEvent::RevisionChanged(r) => Some(r),
                _ => None,
            })
            .collect()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::SheetId;

    #[test]
    fn test_event_collector_filtering() {
        let mut collector = EventCollector::new();

        collector.push(WorkbookEvent::RevisionChanged(RevisionChangedEvent {
            revision: 1,
            previous: 0,
        }));
        collector.push(WorkbookEvent::CellsChanged(CellsChangedEvent {
            revision: 1,
            cells: vec![CellId::new(SheetId(1), 0, 0)],
        }));
        collector.push(WorkbookEvent::BatchApplied(BatchAppliedEvent {
            revision: 1,
            applied: 1,
            total: 1,
            error: None,
        }));

        assert_eq!(collector.len(), 3);
        assert_eq!(collector.batch_applied().len(), 1);
        assert_eq!(collector.cells_changed().len(), 1);
        assert_eq!(collector.revision_changed().len(), 1);
    }
}
